import { useState } from 'react';
import { DEFAULT_CELL_MARGINS, DEFAULT_TABLE_BORDER, TABLE_BORDER_SIDES, tableColumnWeights } from './model';
import type { CellAlignment, Table, TableBorder, TableFormatting } from './model';

type BorderPreset = 'grid' | 'outline' | 'horizontal' | 'none' | 'custom';

function borderPreset(table: Table): BorderPreset {
  if (!table.borders) return 'grid';
  const visible = TABLE_BORDER_SIDES.map(side => {
    const border = table.borders?.[side];
    return !!border && border.style !== 'none';
  });
  if (visible.every(Boolean)) return 'grid';
  if (visible.every(value => !value)) return 'none';
  if (visible.slice(0, 4).every(Boolean) && !visible[4] && !visible[5]) return 'outline';
  if (visible[0] && visible[2] && visible[4] && !visible[1] && !visible[3] && !visible[5]) return 'horizontal';
  return 'custom';
}

function borderSample(table: Table) {
  return TABLE_BORDER_SIDES.map(side => table.borders?.[side])
    .find(border => border && border.style !== 'none') ?? DEFAULT_TABLE_BORDER;
}

const round = (value: number) => Math.round(value * 10) / 10;
const borderSize = (value: number) => Math.min(12, Math.max(.25, value));

export function TableFormatFields({ table, row, column, whole, availableWidth }: {
  table: Table;
  row: number;
  column: number;
  whole: boolean;
  availableWidth: number;
}) {
  const weights = tableColumnWeights(table);
  const percent = 100 * weights[column] / weights.reduce((sum, value) => sum + value, 0);
  const currentWidth = table.width?.unit === 'px' ? table.width.value / availableWidth * 100 : table.width?.value ?? 100;
  const margins = { ...DEFAULT_CELL_MARGINS, ...table.margins };
  const preset = borderPreset(table);
  const border = borderSample(table);
  const [pattern, setPattern] = useState(preset);
  const [shading, setShading] = useState('keep');
  const [equalColumns, setEqualColumns] = useState(false);
  return <div className="rwe-table-format">
    <fieldset>
      <legend>Table</legend>
      <div className="rwe-dialog-grid">
        <label>Table width (%)<input name="table-width" type="number" min="10" max={Math.max(100, round(currentWidth))}
          step=".1" defaultValue={round(currentWidth)} required /></label>
        <label>Table alignment<select name="table-align" aria-label="Table alignment" defaultValue={table.align ?? 'left'}>
          <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
        </select></label>
      </div>
      <div className="rwe-dialog-grid">
        <label>Borders<select name="table-borders" aria-label="Borders" value={pattern}
          onChange={event => setPattern(event.target.value as BorderPreset)}>
          {preset === 'custom' && <option value="custom">Existing borders</option>}
          <option value="grid">All borders</option><option value="outline">Outside borders</option>
          <option value="horizontal">Horizontal borders</option><option value="none">No borders</option>
        </select></label>
        <label>Border style<select name="table-border-style" aria-label="Border style" defaultValue={border.style}
          disabled={pattern === 'none'}>
          <option value="solid">Solid</option><option value="dashed">Dashed</option>
          <option value="dotted">Dotted</option><option value="double">Double</option>
        </select></label>
        <label>Border width (pt)<input name="table-border-width" type="number" min=".25" max="12" step="any"
          defaultValue={borderSize(border.size)} required disabled={pattern === 'none'} /></label>
        <label>Border color<input name="table-border-color" type="color" defaultValue={border.color}
          disabled={pattern === 'none'} /></label>
      </div>
      <label className="rwe-table-format-check"><input name="table-repeat-header" type="checkbox" defaultChecked={!!table.rows[0].header} />
        Repeat header row on each page</label>
    </fieldset>
    <fieldset>
      <legend>Size and spacing</legend>
      <div className="rwe-dialog-grid">
        <label>Column {column + 1} width (%)<input name="table-column-width" type="number" min={weights.length > 1 ? Math.min(2, round(percent)) : 100}
          max={weights.length > 1 ? Math.max(98, round(percent)) : 100} step=".1" defaultValue={round(percent)} required
          disabled={weights.length === 1 || equalColumns} /></label>
        <label>Row {row + 1} minimum height (px)<input name="table-row-height" type="number" min="0"
          max={Math.max(1000, round(table.rows[row].height ?? 0))} step=".1" defaultValue={round(table.rows[row].height ?? 0)} required /></label>
        <label>Horizontal cell padding (px)<input name="table-padding-x" type="number" min="0"
          max={Math.max(96, round(margins.left))} step=".1" defaultValue={round(margins.left)} required /></label>
        <label>Vertical cell padding (px)<input name="table-padding-y" type="number" min="0"
          max={Math.max(96, round(margins.top))} step=".1" defaultValue={round(margins.top)} required /></label>
      </div>
      <label className="rwe-table-format-check"><input name="table-equal-columns" type="checkbox" disabled={weights.length === 1}
        checked={equalColumns} onChange={event => setEqualColumns(event.target.checked)} />
        Distribute columns evenly</label>
    </fieldset>
    <fieldset>
      <legend>Cells</legend>
      <div className="rwe-dialog-grid">
        <label>Apply cell formatting to<select name="table-cell-scope" aria-label="Apply cell formatting to" defaultValue={whole ? 'table' : 'cell'}>
          <option value="cell">Current cell</option><option value="row">Current row</option>
          <option value="column">Current column</option><option value="table">Entire table</option>
        </select></label>
        <label>Vertical alignment<select name="table-cell-align" aria-label="Vertical alignment" defaultValue="">
          <option value="">Keep current alignment</option><option value="top">Top</option>
          <option value="center">Middle</option><option value="bottom">Bottom</option>
        </select></label>
        <label>Shading<select name="table-shading" aria-label="Shading" value={shading} onChange={event => setShading(event.target.value)}>
          <option value="keep">Keep current shading</option><option value="color">Set a color</option><option value="clear">Clear shading</option>
        </select></label>
        <label>Shading color<input name="table-shading-color" type="color" defaultValue={table.rows[row].cells[column].fill ?? '#e8effc'}
          disabled={shading !== 'color'} /></label>
      </div>
    </fieldset>
  </div>;
}

export function readTableFormatting(data: FormData, table: Table, row: number, column: number, availableWidth: number): TableFormatting {
  const result: TableFormatting = { table: {} };
  const props = result.table!;
  const number = (key: string) => Number(data.get(key));
  const currentWidth = table.width?.unit === 'px' ? table.width.value / availableWidth * 100 : table.width?.value ?? 100;
  const width = number('table-width');
  if (width !== round(currentWidth)) props.width = { unit: 'percent', value: width };
  const align = data.get('table-align') as Table['align'];
  if (align !== (table.align ?? 'left')) props.align = align;

  const preset = data.get('table-borders') as BorderPreset;
  const originalPreset = borderPreset(table);
  const border = borderSample(table);
  const style = data.get('table-border-style') as TableBorder['style'] | null;
  const color = data.get('table-border-color') as string | null;
  const size = Math.round(borderSize(number('table-border-width')) * 8) / 8;
  const styleChanged = style !== null && style !== border.style;
  const colorChanged = color !== null && color !== border.color;
  const sizeChanged = size !== borderSize(border.size);
  if (preset !== originalPreset) {
    props.borders = Object.fromEntries(TABLE_BORDER_SIDES.map(side => {
      const on = preset === 'grid' || (preset === 'outline' && side !== 'insideH' && side !== 'insideV') ||
        (preset === 'horizontal' && (side === 'top' || side === 'bottom' || side === 'insideH'));
      return [side, { style: on ? style ?? 'solid' : 'none', color: color ?? border.color, size }];
    }));
  } else if (preset !== 'none' && (styleChanged || colorChanged || sizeChanged)) {
    props.borders = Object.fromEntries(TABLE_BORDER_SIDES.flatMap(side => {
      const existing = table.borders ? table.borders[side] : DEFAULT_TABLE_BORDER;
      if (!existing) return [];
      if (existing.style === 'none') return [[side, existing]];
      return [[side, {
        style: styleChanged ? style : existing.style,
        color: colorChanged ? color : existing.color,
        size: sizeChanged ? size : existing.size,
      }]];
    }));
  }

  const margins = { ...DEFAULT_CELL_MARGINS, ...table.margins };
  const x = number('table-padding-x'), y = number('table-padding-y');
  if (x !== round(margins.left) || y !== round(margins.top)) props.margins = {
    ...margins, ...(x !== round(margins.left) ? { left: x, right: x } : {}), ...(y !== round(margins.top) ? { top: y, bottom: y } : {}),
  };

  const weights = tableColumnWeights(table);
  const percent = weights[column] / weights.reduce((sum, value) => sum + value, 0) * 100;
  const newPercent = number('table-column-width');
  if (data.has('table-equal-columns')) result.columnPercent = 'equal';
  else if (weights.length > 1 && newPercent !== round(percent)) result.columnPercent = newPercent;
  const height = number('table-row-height');
  if (height !== round(table.rows[row].height ?? 0)) result.rowHeight = height;
  const repeat = data.has('table-repeat-header');
  if (repeat !== !!table.rows[0].header) result.repeatHeader = repeat;

  const shading = data.get('table-shading');
  const vertical = data.get('table-cell-align') as CellAlignment | '';
  if (shading !== 'keep' || vertical) result.cells = {
    scope: data.get('table-cell-scope') as NonNullable<TableFormatting['cells']>['scope'],
    ...(shading !== 'keep' ? { fill: shading === 'color' ? String(data.get('table-shading-color')) : undefined } : {}),
    ...(vertical ? { vertical } : {}),
  };
  return result;
}
