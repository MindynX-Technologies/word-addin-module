import { useEffect, useMemo, useRef, useState } from 'react';
import { FilePlus2, FolderOpen, Save, Trash2 } from 'lucide-react';
import { fieldsIn, fillTemplate, prepareTemplate, templateFileName } from './templates';
import { listTemplates, removeTemplate, saveTemplate } from './template-library';
import type { StoredTemplate } from './template-library';
import type { ImportedDocx } from './docx-read';
import type { DocModel } from './model';

type Props = {
  fileName: string;
  exportDocument: () => Promise<File>;
  create: (model: DocModel, template: ImportedDocx, fileName: string) => void;
  allowSave: boolean;
};

export function TemplateDialog({ fileName, exportDocument, create, allowSave }: Props) {
  const [templates, setTemplates] = useState<StoredTemplate[]>([]);
  const [selected, setSelected] = useState<{ record: StoredTemplate; document: ImportedDocx }>();
  const [saveName, setSaveName] = useState(fileName.replace(/\.do[ct]x$/i, ''));
  const [newName, setNewName] = useState('');
  const [values, setValues] = useState<Record<string, string>>({});
  const [busy, setBusy] = useState(true);
  const [error, setError] = useState('');
  const [remove, setRemove] = useState<string>();
  const upload = useRef<HTMLInputElement>(null);

  useEffect(() => {
    let active = true;
    listTemplates().then(records => {
      if (active) setTemplates(records);
    }, () => {
      if (active) setError('The browser could not open the template library.');
    }).finally(() => { if (active) setBusy(false); });
    return () => { active = false; };
  }, []);

  async function perform(action: () => Promise<void>) {
    setBusy(true);
    setError('');
    try { await action(); }
    catch (cause) { setError(cause instanceof Error ? cause.message : 'The template operation failed.'); }
    finally { setBusy(false); }
  }

  async function choose(record: StoredTemplate) {
    const document = await prepareTemplate(record.file);
    setSelected({ record, document });
    setNewName(record.name);
    setValues({});
  }

  async function save(file: File, name: string) {
    const document = await prepareTemplate(file);
    const record = await saveTemplate(name, file);
    setTemplates(previous => [record, ...previous]);
    setSelected({ record, document });
    setNewName(name);
    setValues({});
  }

  const fields = useMemo(() => selected ? fieldsIn(selected.document.model) : [], [selected]);
  return <div className="rwe-templates">
    <p>Mark fillable text in a document with <code>{'{{client_name}}'}</code>, then save it as a template.</p>
    {error && <p className="rwe-template-error" role="alert">{error}</p>}
    <section aria-label="Save and import templates">
      <form className="rwe-template-controls" onSubmit={event => {
        event.preventDefault();
        const name = saveName.trim();
        if (!name || !allowSave) return;
        void perform(async () => save(await exportDocument(), name));
      }}>
        <label>Template name<input value={saveName} onChange={event => setSaveName(event.target.value)} maxLength={100} required /></label>
        <button className="rwe-primary" type="submit" disabled={busy || !allowSave || !saveName.trim()}><Save size={15} /> Save current document</button>
        <button type="button" disabled={busy} onClick={() => upload.current?.click()}><FolderOpen size={15} /> Import Word file</button>
        <input ref={upload} type="file" accept=".docx,.dotx" hidden onChange={event => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file) void perform(() => save(file, file.name.replace(/\.do[ct]x$/i, '').slice(0, 100)));
        }} />
      </form>
    </section>
    <div className="rwe-template-library">
      <section aria-label="Saved templates">
        <h3>Saved templates</h3>
        {templates.length ? <ul>{templates.map(record => <li key={record.id}>
          <button type="button" className="rwe-template-item" aria-pressed={selected?.record.id === record.id} disabled={busy}
            onClick={() => void perform(() => choose(record))}>{record.name}</button>
          <button type="button" className="rwe-template-remove" aria-label={'Delete ' + record.name} disabled={busy}
            onClick={() => setRemove(record.id)}><Trash2 size={14} /></button>
          {remove === record.id && <div className="rwe-template-confirm">
            <span>Delete this template?</span>
            <button type="button" onClick={() => setRemove(undefined)}>Cancel</button>
            <button type="button" onClick={() => void perform(async () => {
              await removeTemplate(record.id);
              setTemplates(previous => previous.filter(item => item.id !== record.id));
              if (selected?.record.id === record.id) setSelected(undefined);
              setRemove(undefined);
            })}>Delete</button>
          </div>}
        </li>)}</ul> : <p>{busy ? 'Loading templates…' : 'No templates saved yet.'}</p>}
      </section>
      {selected && <form aria-label="Create from template" onSubmit={event => {
        event.preventDefault();
        void perform(async () => create(fillTemplate(selected.document.model, values), selected.document, templateFileName(newName)));
      }}>
        <h3>{selected.record.name}</h3>
        <label>Document name<input value={newName} onChange={event => setNewName(event.target.value)} maxLength={120} required /></label>
        {fields.map(field => <label key={field.name}>{field.name}
          <textarea rows={2} value={values[field.name] ?? ''} onChange={event => setValues(previous => ({
            ...previous, [field.name]: event.target.value,
          }))} maxLength={20_000} required={!(field.name in values)} />
        </label>)}
        <button className="rwe-primary" type="submit" disabled={busy || !newName.trim()}><FilePlus2 size={15} /> Create document</button>
      </form>}
    </div>
  </div>;
}
