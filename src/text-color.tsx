import { ChevronDown } from 'lucide-react';
import { useEffect, useId, useRef, useState } from 'react';

const COLORS = [
  ['Black', '#000000'], ['Charcoal', '#24272d'], ['Slate', '#404b5e'], ['Gray', '#808080'],
  ['Silver', '#a6a6a6'], ['Light gray', '#d9d9d9'], ['Ivory', '#fff2cc'], ['White', '#ffffff'],
  ['Dark red', '#a42d30'], ['Red', '#c00000'], ['Orange', '#ed7d31'], ['Gold', '#bf9000'],
  ['Green', '#548235'], ['Teal', '#008064'], ['Blue', '#1663d6'], ['Purple', '#693ba7'],
  ['Rose', '#e6b8b7'], ['Pink', '#f4b6d1'], ['Peach', '#f8cbad'], ['Yellow', '#ffd966'],
  ['Light green', '#a9d18e'], ['Mint', '#8fd8cc'], ['Light blue', '#9dc3e6'], ['Lavender', '#c9b3e6'],
] as const;
const COLUMNS = 8;

export function TextColor({ value = '#24272d', disabled, onChange }: {
  value?: string;
  disabled: boolean;
  onChange: (color: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [custom, setCustom] = useState(value);
  const selected = COLORS.findIndex(([, color]) => color === value.toLowerCase());
  const [active, setActive] = useState(0);
  const root = useRef<HTMLDivElement>(null);
  const trigger = useRef<HTMLButtonElement>(null);
  const palette = useRef<HTMLDivElement>(null);
  const keyboardOpened = useRef(false);
  const id = useId();

  useEffect(() => {
    if (!open || disabled) return;
    const document = root.current?.ownerDocument;
    function dismiss(event: PointerEvent) {
      if (!event.composedPath().includes(root.current!)) setOpen(false);
    }
    function escape(event: KeyboardEvent) {
      if (event.key !== 'Escape' || event.defaultPrevented) return;
      event.preventDefault();
      event.stopPropagation();
      setOpen(false);
      if (root.current?.contains(document?.activeElement ?? null)) trigger.current?.focus();
    }
    document?.addEventListener('pointerdown', dismiss);
    document?.addEventListener('keydown', escape, true);
    if (keyboardOpened.current) palette.current?.querySelector<HTMLButtonElement>('[tabindex="0"]')?.focus();
    return () => {
      document?.removeEventListener('pointerdown', dismiss);
      document?.removeEventListener('keydown', escape, true);
    };
  }, [open, disabled]);

  function pick(color: string) {
    setOpen(false);
    onChange(color);
  }

  return <div ref={root} className="rwe-color" onBlur={event => {
    if (event.relatedTarget && !event.currentTarget.contains(event.relatedTarget)) setOpen(false);
  }}>
    <button ref={trigger} className="rwe-tool rwe-color-trigger" type="button" aria-label="Text color"
      aria-expanded={open && !disabled} aria-controls={open && !disabled ? id : undefined} title="Text color"
      disabled={disabled} onMouseDown={event => event.preventDefault()} onClick={event => {
        if (!open) {
          setCustom(value);
          setActive(Math.max(0, selected));
          keyboardOpened.current = event.detail === 0;
        }
        setOpen(!open);
      }}>
      <span aria-hidden="true" style={{ borderColor: value }}>A</span><ChevronDown aria-hidden="true" />
    </button>
    {open && !disabled && <div id={id} className="rwe-color-palette" role="group" aria-label="Text color palette">
      <div ref={palette} className="rwe-color-swatches" onKeyDown={event => {
        const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -COLUMNS, ArrowDown: COLUMNS };
        const offset = offsets[event.key];
        if (offset === undefined) return;
        event.preventDefault();
        const next = Math.max(0, Math.min(COLORS.length - 1, active + offset));
        setActive(next);
        event.currentTarget.querySelectorAll('button')[next]?.focus();
      }}>
        {COLORS.map(([name, color], index) => <button key={color} type="button" title={name}
          className="rwe-color-swatch" aria-label={name} aria-pressed={index === selected}
          style={{ backgroundColor: color }} tabIndex={index === active ? 0 : -1}
          onMouseDown={event => event.preventDefault()} onClick={() => pick(color)} />)}
      </div>
      <div className="rwe-color-custom">
        <label>Custom color <input type="color" value={custom} onChange={event => setCustom(event.target.value)} /></label>
        <button type="button" onMouseDown={event => event.preventDefault()} onClick={() => pick(custom)}>Apply</button>
      </div>
    </div>}
  </div>;
}
