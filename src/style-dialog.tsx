import { DEFAULT_FORMAT } from './model';
import type { Alignment, NamedStyle } from './model';
import { useState } from 'react';

export const FONTS = ['Arial', 'Aptos', 'Calibri', 'Cambria', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New'];

export function StyleFields({ style }: { style: NamedStyle }) {
  const format = { ...DEFAULT_FORMAT, ...style.format };
  const paragraph = style.paragraph;
  const [lineRule, setLineRule] = useState<'auto' | 'exact' | 'atLeast'>(paragraph.lineSpacing?.rule ?? 'auto');
  return <div className="rwe-style-fields">
    <label>Style name<input name="style-name" defaultValue={style.name} maxLength={80} required autoFocus /></label>
    <div className="rwe-dialog-grid">
      <label>Font family<input name="style-font" defaultValue={format.font} list="rwe-style-fonts" maxLength={80} required /></label>
      <datalist id="rwe-style-fonts">{FONTS.map(font => <option key={font} value={font} />)}</datalist>
      <label>Font size (pt)<input name="style-size" type="number" min="1" max="400" step=".5" defaultValue={format.size} required /></label>
    </div>
    <div className="rwe-style-toggles">
      <label><input name="style-bold" type="checkbox" defaultChecked={format.bold} /> Bold</label>
      <label><input name="style-italic" type="checkbox" defaultChecked={format.italic} /> Italic</label>
      <label><input name="style-underline" type="checkbox" defaultChecked={format.underline} /> Underline</label>
    </div>
    <div className="rwe-dialog-grid">
      <label>Text color<input name="style-color" type="color" defaultValue={format.color} /></label>
      <label>Alignment<select aria-label="Alignment" name="style-alignment" defaultValue={paragraph.alignment ?? 'left'}>
        <option value="left">Left</option><option value="center">Center</option>
        <option value="right">Right</option><option value="justify">Justified</option>
      </select></label>
      <label>Space before (pt)<input name="style-before" type="number" min="0" max="200" step="any" defaultValue={points(paragraph.before, 0)} required /></label>
      <label>Space after (pt)<input name="style-after" type="number" min="0" max="200" step="any" defaultValue={points(paragraph.after, 4)} required /></label>
      <label>Line spacing rule<select aria-label="Line spacing rule" name="style-line-rule" value={lineRule}
        onChange={event => setLineRule(event.target.value as typeof lineRule)}>
        <option value="auto">Multiple</option><option value="exact">Exactly (pt)</option><option value="atLeast">At least (pt)</option>
      </select></label>
      <label>Line spacing<input key={lineRule} name="style-line" type="number" min={lineRule === 'auto' ? .8 : 1}
        max={lineRule === 'auto' ? 6 : 400} step="any"
        defaultValue={lineRule === 'auto' ? paragraph.lineHeight ?? 1.32 : (paragraph.lineSpacing?.value ?? 16) * .75} required /></label>
      <label>Outline level<select aria-label="Outline level" name="style-outline" defaultValue={style.heading ?? 0}>
        <option value="0">Body text</option>
        {Array.from({ length: 9 }, (_, index) => <option key={index} value={index + 1}>Heading {index + 1}</option>)}
      </select></label>
      <label>Left indent (pt)<input name="style-indent" type="number" min="0" max="200" step="any" defaultValue={points(paragraph.indent, 0)} required /></label>
      <label>First-line indent (pt)<input name="style-first-line" type="number" min="-200" max="200" step="any" defaultValue={points(paragraph.firstLine, 0)} required /></label>
    </div>
  </div>;
}

const points = (value: number | undefined, fallback: number) => Number(((value ?? fallback) * .75).toFixed(3));

export function readStyleFields(data: FormData, base: NamedStyle, id: string, create: boolean): NamedStyle {
  const name = String(data.get('style-name')).trim();
  const font = String(data.get('style-font')).trim();
  const number = (key: string) => Number(data.get('style-' + key));
  const px = (key: string) => number(key) * 4 / 3;
  const level = number('outline');
  const lineRule = String(data.get('style-line-rule') ?? 'auto');
  const values = {
    paragraph: {
      alignment: data.get('style-alignment') as Alignment,
      before: px('before'), after: px('after'), lineHeight: lineRule === 'auto' ? number('line') : undefined,
      lineSpacing: lineRule === 'exact' || lineRule === 'atLeast' ? { rule: lineRule as 'exact' | 'atLeast', value: px('line') } : undefined,
      indent: px('indent'), firstLine: px('first-line'),
    },
    format: {
      font, size: number('size'), color: String(data.get('style-color')),
      bold: data.has('style-bold'), italic: data.has('style-italic'), underline: data.has('style-underline'),
    },
    outline: level ? level - 1 : 9,
  };
  const format = { ...base.format, ...values.format };
  const paragraph = { ...base.paragraph, ...values.paragraph };
  const local = create ? { ...values, paragraph, format } : {
    paragraph: { ...base.local?.paragraph, ...values.paragraph },
    format: { ...base.local?.format, ...values.format }, outline: values.outline,
  };
  return { ...base, id, name, type: 'paragraph', format, paragraph, heading: level || undefined,
    basedOn: create ? base.id : base.basedOn, local };
}
