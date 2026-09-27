import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../types.js';
import { viewportPreviewBlocker } from './worksheet-preview-eligibility.js';

const viewport = { row: 1, col: 1, rows: 30, cols: 12 };

function sheet(conditionalFormats: Worksheet['conditionalFormats']): Worksheet {
  return { conditionalFormats, charts: [], images: [], shapeGroups: [],
    sparklineGroups: [] } as unknown as Worksheet;
}

describe('first viewport dependency gate', () => {
  it('waits for whole-range statistics only when their range intersects the viewport', () => {
    const rule = { type: 'top10' as const, top: true, percent: false, rank: 1,
      dxfId: 0, priority: 1 };
    expect(viewportPreviewBlocker(sheet([{ sqref: [{ top: 1, left: 1, bottom: 70000, right: 1 }],
      rules: [rule] }]), viewport, 128)).toBe('conditional-format-range');
    expect(viewportPreviewBlocker(sheet([{ sqref: [{ top: 100, left: 1, bottom: 70000, right: 1 }],
      rules: [rule] }]), viewport, 128)).toBeNull();
    expect(viewportPreviewBlocker(sheet([{ sqref: [{ top: 1, left: 1, bottom: 30, right: 1 }],
      rules: [rule] }]), viewport, 128)).toBeNull();
    expect(viewportPreviewBlocker(sheet([{ sqref: [
      { top: 1, left: 1, bottom: 30, right: 1 },
      { top: 200, left: 1, bottom: 70000, right: 1 },
    ], rules: [rule] }]), viewport, 128)).toBe('conditional-format-range');
  });

  it('permits a row-local comparison and a drawing anchored below the viewport', () => {
    const ws = sheet([{ sqref: [{ top: 1, left: 1, bottom: 70000, right: 1 }],
      rules: [{ type: 'cellIs', operator: 'greaterThan', formulas: ['10'], dxfId: 0, priority: 1 }] }]);
    ws.charts = [{ fromRow: 500, fromCol: 1, toRow: 520, toCol: 12 }] as Worksheet['charts'];
    expect(viewportPreviewBlocker(ws, viewport, 128)).toBeNull();
    ws.charts = [{ fromRow: 0, fromCol: 1, toRow: 15, toCol: 12 }] as Worksheet['charts'];
    expect(viewportPreviewBlocker(ws, viewport, 128)).toBe('drawing-dependency');
  });
});
