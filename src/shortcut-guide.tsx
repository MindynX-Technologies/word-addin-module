import { useMemo, useState } from 'react';
import { shortcutGroups } from './shortcuts';

export function ShortcutGuide({ mac }: { mac: boolean }) {
  const [query, setQuery] = useState('');
  const groups = useMemo(() => shortcutGroups(mac), [mac]);
  const terms = query.toLowerCase().trim().split(/\s+/);
  const results = groups.map(group => ({
    ...group,
    entries: group.entries.filter(({ label, keys }) =>
      terms.every(term => [group.group, label, ...keys].join(' ').toLowerCase().includes(term))),
  })).filter(group => group.entries.length);

  return <>
    <label>Search shortcuts<input autoFocus type="search" value={query} onChange={event => setQuery(event.target.value)}
      placeholder="Try bold, headings, or review" /></label>
    <p>Shortcuts apply when the editor has focus. Your browser or system may reserve some keys.</p>
    <div className="rwe-shortcut-list">
      {results.map(({ group, entries }) => <section key={group} aria-label={group}>
        <h3>{group}</h3>
        <dl>{entries.map(entry => <div key={entry.id}>
          <dt>{entry.label}</dt>
          <dd>{entry.keys.map(keys => <kbd key={keys}>{keys}</kbd>)}</dd>
        </div>)}</dl>
      </section>)}
      {!results.length && <p role="status">No matching shortcuts.</p>}
    </div>
  </>;
}
